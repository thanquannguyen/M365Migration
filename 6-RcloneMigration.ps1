<#
.SYNOPSIS
Reliable cross-tenant OneDrive migration orchestrator using rclone and Entra ID app-only authentication.

.DESCRIPTION
Resolves source/destination OneDrive drive IDs with Microsoft Graph, creates a temporary
rclone configuration that uses rclone's native OAuth client-credentials flow, and copies
OneDrive content between tenants. Existing destination files are always skipped. After each successful user migration, old OneDrive versions on that destination drive are cleaned only when -EnableVersionCleanup is supplied (opt-in: cleanup permanently deletes old versions). Use -CleanupOnly together with -EnableVersionCleanup to clean all mapped destination users without copying.

Unlike a configuration containing a pre-generated one-hour access token, client-credentials
mode lets rclone obtain fresh tokens during long-running transfers.

Application credentials are configured directly in this script, matching the original
migration script style. Client secrets are written directly to the short-lived temporary
rclone configuration and the file is deleted after each user. The copy command also enables
--onedrive-av-override by default so OneDrive items flagged by Microsoft AV are force-downloaded
through rclone's supported override path instead of being blocked. On low-memory Windows hosts, the script also applies conservative rclone/Go memory limits, checks page-file status, and retries confirmed OOM crashes with progressively lower concurrency.

Examples:
  # Normal migration: requests up to 3 users in parallel. On low-memory hosts the script
  # automatically caps account-level concurrency (1 user below 2 GiB RAM, 2 below 4 GiB).
  .\6-RcloneMigration.ps1 -MappingPath '.\Output\UserMapping.csv'

  # Override account-level concurrency.
  .\6-RcloneMigration.ps1 -ParallelUsers 2

  # Cleanup old versions for every mapped destination user without copying anything.
  .\6-RcloneMigration.ps1 -CleanupOnly -EnableVersionCleanup

  # Preview copy + cleanup actions without changing data.
  .\6-RcloneMigration.ps1 -DryRun

  # Delete old destination versions after each successful migration (opt-in).
  # Without this switch, destination version history is preserved.
  .\6-RcloneMigration.ps1 -EnableVersionCleanup

.NOTES
The source application needs Microsoft Graph application permissions sufficient to read the
source OneDrive. The destination application needs application permissions sufficient to write
to the destination OneDrive. Admin consent must be granted.
#>

[CmdletBinding()]
param (
    [string]$ConfigPath = '.\MigrationConfig.psd1',
    [string]$MappingPath = '.\Output\UserMapping.csv',
    [string]$UserFilter,
    [string]$RclonePath = 'rclone',

    # Number of destination accounts requested simultaneously. Default remains 3,
    # but the low-memory guard automatically caps this to 1 on machines below 2 GiB RAM
    # and to 2 below 4 GiB. Use -ForceParallelUsers to bypass that safety cap.
    [ValidateRange(1, 12)]
    [int]$ParallelUsers = 3,

    [switch]$ForceParallelUsers,

    # Internal flag used by the parent process when launching one-user workers.
    [switch]$InternalWorker,

    # Low-memory defaults suitable for the 1 GiB / 2 vCPU migration VPS.
    [ValidateRange(1, 64)]
    [int]$Transfers = 2,

    [ValidateRange(1, 128)]
    [int]$Checkers = 4,

    [ValidatePattern('^[0-9]+(Ki|Mi|Gi|K|M|G)?$')]
    [string]$ChunkSize = '5Mi',

    # Keep the queued object backlog small so large drives do not consume excessive RAM.
    [ValidateRange(100, 100000)]
    [int]$MaxBacklog = 1000,

    # Soft Go runtime memory target inherited by each rclone process.
    [ValidateRange(128, 8192)]
    [int]$GoMemoryLimitMiB = 512,

    # More aggressive Go garbage collection on memory-constrained hosts.
    # rclone's FAQ recommends GOGC=20 when reducing memory use is more important than CPU.
    [ValidateRange(5, 100)]
    [int]$GoGCPercent = 20,

    [ValidateRange(1, 20)]
    [int]$RcloneRetries = 5,

    [ValidateRange(1, 20)]
    [int]$MaxAttempts = 3,

    [ValidateRange(1, 300)]
    [int]$RetryDelaySeconds = 15,

    [switch]$UseChecksum,

    [switch]$DryRun,
    [switch]$StopOnError,

    # Cleanup of old OneDrive versions is OPT-IN: it permanently deletes old
    # destination versions (deleted versions cannot be recovered), so it only
    # runs when -EnableVersionCleanup is supplied explicitly. By default,
    # destination version history is preserved. -CleanupOnly requires this switch.
    [switch]$EnableVersionCleanup,

    # Skip copying and only run version cleanup for every mapped destination user.
    # Requires -EnableVersionCleanup.
    [switch]$CleanupOnly,

    [ValidateRange(1, 128)]
    [int]$CleanupCheckers = 4
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

# Shared safety helpers: required config values throw instead of falling back
# to a hard-coded tenant or credential.
. "$PSScriptRoot\MigrationHelpers.ps1"

# --- Credentials ---
$config = @{}
if (Test-Path $ConfigPath) {
    try { $config = Import-PowerShellDataFile -Path $ConfigPath } catch {}
}

# All tenant/credential values are required: a missing key or a leftover
# template placeholder throws here instead of failing obscurely later.
$SrcTenantId = Get-RequiredConfigValue -Value $config.SourceTenantDomain -Name 'SourceTenantDomain'
$SrcClientId = Get-RequiredConfigValue -Value $config.SourceClientId -Name 'SourceClientId'
$SrcClientSecret = Get-RequiredConfigValue -Value $config.SourceClientSecret -Name 'SourceClientSecret'

$DstTenantId = Get-RequiredConfigValue -Value $config.DestTenantDomain -Name 'DestTenantDomain'
$DstClientId = Get-RequiredConfigValue -Value $config.DestClientId -Name 'DestClientId'
$DstClientSecret = Get-RequiredConfigValue -Value $config.DestClientSecret -Name 'DestClientSecret'

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
if ([string]::IsNullOrWhiteSpace($ScriptDir)) {
    $ScriptDir = (Get-Location).Path
}

# Resolve relative paths against the script directory without changing the caller's location.
if (-not [System.IO.Path]::IsPathRooted($MappingPath)) {
    $MappingPath = Join-Path -Path $ScriptDir -ChildPath $MappingPath
}

$LogDir = Join-Path -Path $ScriptDir -ChildPath 'Logs'
if (-not (Test-Path -LiteralPath $LogDir)) {
    New-Item -ItemType Directory -Path $LogDir -Force | Out-Null
}

$RunTimestamp = "$(Get-Date -Format 'yyyyMMdd_HHmmss_fff')_${PID}"
$LogPath = Join-Path -Path $LogDir -ChildPath "RcloneMigration_${RunTimestamp}.log"
$SummaryPath = Join-Path -Path $LogDir -ChildPath "RcloneMigration_${RunTimestamp}_Summary.csv"

function Write-Log {
    param (
        [Parameter(Mandatory = $true)][string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR', 'SUCCESS')][string]$Type = 'INFO'
    )

    $logMessage = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$Type] $Message"
    $color = 'Cyan'
    if ($Type -eq 'ERROR') {
        $color = 'Red'
    }
    elseif ($Type -eq 'WARN') {
        $color = 'Yellow'
    }
    elseif ($Type -eq 'SUCCESS') {
        $color = 'Green'
    }

    Write-Host $logMessage -ForegroundColor $color
    Add-Content -LiteralPath $LogPath -Value $logMessage -Encoding UTF8 -ErrorAction SilentlyContinue
}

function Resolve-RcloneExecutable {
    param ([Parameter(Mandatory = $true)][string]$Path)

    if (Test-Path -LiteralPath $Path -PathType Leaf) {
        return (Resolve-Path -LiteralPath $Path).Path
    }

    $command = Get-Command -Name $Path -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $command) {
        throw "rclone executable not found: $Path"
    }

    return $command.Source
}

function Test-RcloneClientCredentialsSupport {
    param ([Parameter(Mandatory = $true)][string]$Executable)

    $versionOutput = (& $Executable version 2>&1 | Out-String).Trim()
    if ($LASTEXITCODE -ne 0) {
        throw "Unable to execute rclone: $versionOutput"
    }

    $firstLine = ($versionOutput -split "`r?`n" | Select-Object -First 1)
    Write-Log "Using $firstLine"

    $helpOutput = (& $Executable help backend onedrive 2>&1 | Out-String)
    if ($LASTEXITCODE -ne 0 -or $helpOutput -notmatch 'client_credentials') {
        throw 'This rclone build does not appear to support OneDrive client_credentials. Update rclone before running this migration.'
    }
}

function Get-GraphAccessToken {
    param (
        [Parameter(Mandatory = $true)][string]$TenantId,
        [Parameter(Mandatory = $true)][string]$ClientId,
        [Parameter(Mandatory = $true)][string]$ClientSecret,
        [ValidateRange(1, 10)][int]$Attempts = 4
    )

    $uri = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    $body = @{
        client_id     = $ClientId
        client_secret = $ClientSecret
        scope         = 'https://graph.microsoft.com/.default'
        grant_type    = 'client_credentials'
    }

    for ($attempt = 1; $attempt -le $Attempts; $attempt++) {
        try {
            $response = Invoke-RestMethod -Uri $uri -Method Post -Body $body -ContentType 'application/x-www-form-urlencoded' -TimeoutSec 60 -ErrorAction Stop
            if ([string]::IsNullOrWhiteSpace($response.access_token)) {
                throw 'Token endpoint returned no access_token.'
            }
            return $response.access_token
        }
        catch {
            if ($attempt -ge $Attempts) {
                throw
            }

            $delay = [Math]::Min(30, [Math]::Pow(2, $attempt))
            Write-Log "Token request failed (attempt $attempt/$Attempts). Retrying in $delay seconds. Error: $($_.Exception.Message)" -Type WARN
            Start-Sleep -Seconds $delay
        }
    }
}

function Get-OneDriveInfo {
    param (
        [Parameter(Mandatory = $true)][string]$Token,
        [Parameter(Mandatory = $true)][string]$UserId,
        [ValidateRange(1, 10)][int]$Attempts = 4
    )

    $escapedUserId = [System.Uri]::EscapeDataString($UserId)
    $uri = "https://graph.microsoft.com/v1.0/users/$escapedUserId/drive?`$select=id,driveType,name,webUrl"
    $headers = @{ Authorization = "Bearer $Token" }

    for ($attempt = 1; $attempt -le $Attempts; $attempt++) {
        try {
            $drive = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get -TimeoutSec 60 -ErrorAction Stop
            if ([string]::IsNullOrWhiteSpace($drive.id)) {
                throw "Graph returned no drive ID for $UserId."
            }
            return $drive
        }
        catch {
            if ($attempt -ge $Attempts) {
                throw
            }

            $delay = [Math]::Min(30, [Math]::Pow(2, $attempt))
            Write-Log "Drive lookup failed for $UserId (attempt $attempt/$Attempts). Retrying in $delay seconds. Error: $($_.Exception.Message)" -Type WARN
            Start-Sleep -Seconds $delay
        }
    }
}

function New-RcloneConfig {
    param (
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$SourceDriveId,
        [Parameter(Mandatory = $true)][string]$DestinationDriveId,
        [Parameter(Mandatory = $true)][string]$SourceTenantId,
        [Parameter(Mandatory = $true)][string]$SourceClientId,
        [Parameter(Mandatory = $true)][string]$SourceClientSecret,
        [Parameter(Mandatory = $true)][string]$DestinationTenantId,
        [Parameter(Mandatory = $true)][string]$DestinationClientId,
        [Parameter(Mandatory = $true)][string]$DestinationClientSecret
    )

    $content = @"
[src]
type = onedrive
client_id = $SourceClientId
client_secret = $SourceClientSecret
client_credentials = true
tenant = $SourceTenantId
drive_id = $SourceDriveId
drive_type = business

[dst]
type = onedrive
client_id = $DestinationClientId
client_secret = $DestinationClientSecret
client_credentials = true
tenant = $DestinationTenantId
drive_id = $DestinationDriveId
drive_type = business
"@

    $parentDir = [System.IO.Path]::GetDirectoryName($Path)
    if (-not [string]::IsNullOrWhiteSpace($parentDir) -and -not (Test-Path -LiteralPath $parentDir)) {
        New-Item -ItemType Directory -Path $parentDir -Force | Out-Null
    }
    $utf8NoBom = New-Object -TypeName System.Text.UTF8Encoding -ArgumentList $false
    [System.IO.File]::WriteAllText($Path, $content, $utf8NoBom)
}

function New-RcloneDestinationConfig {
    param (
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$DestinationDriveId,
        [Parameter(Mandatory = $true)][string]$DestinationTenantId,
        [Parameter(Mandatory = $true)][string]$DestinationClientId,
        [Parameter(Mandatory = $true)][string]$DestinationClientSecret
    )

    $content = @"
[dst]
type = onedrive
client_id = $DestinationClientId
client_secret = $DestinationClientSecret
client_credentials = true
tenant = $DestinationTenantId
drive_id = $DestinationDriveId
drive_type = business
"@

    $parentDir = [System.IO.Path]::GetDirectoryName($Path)
    if (-not [string]::IsNullOrWhiteSpace($parentDir) -and -not (Test-Path -LiteralPath $parentDir)) {
        New-Item -ItemType Directory -Path $parentDir -Force | Out-Null
    }
    $utf8NoBom = New-Object -TypeName System.Text.UTF8Encoding -ArgumentList $false
    [System.IO.File]::WriteAllText($Path, $content, $utf8NoBom)
}

function Get-SafeFileNamePart {
    param ([Parameter(Mandatory = $true)][string]$Value)
    return ($Value -replace '[^A-Za-z0-9._-]', '_')
}

function Get-TotalPhysicalMemoryMiB {
    try {
        $computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop
        if ($computerSystem.TotalPhysicalMemory) {
            return [int][Math]::Round(([double]$computerSystem.TotalPhysicalMemory / 1MB))
        }
    }
    catch {
        try {
            $computerSystem = Get-WmiObject -Class Win32_ComputerSystem -ErrorAction Stop
            if ($computerSystem.TotalPhysicalMemory) {
                return [int][Math]::Round(([double]$computerSystem.TotalPhysicalMemory / 1MB))
            }
        }
        catch {
            return $null
        }
    }

    return $null
}

function Get-PageFileInfo {
    try {
        $computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop
        $pageFiles = @(Get-CimInstance -ClassName Win32_PageFileUsage -ErrorAction Stop)

        $allocatedMiB = 0
        $currentUsageMiB = 0
        $peakUsageMiB = 0
        foreach ($pageFile in $pageFiles) {
            $allocatedMiB += [int]$pageFile.AllocatedBaseSize
            $currentUsageMiB += [int]$pageFile.CurrentUsage
            $peakUsageMiB += [int]$pageFile.PeakUsage
        }

        return [pscustomobject]@{
            AutomaticManaged = [bool]$computerSystem.AutomaticManagedPagefile
            FileCount        = $pageFiles.Count
            AllocatedMiB     = $allocatedMiB
            CurrentUsageMiB  = $currentUsageMiB
            PeakUsageMiB     = $peakUsageMiB
        }
    }
    catch {
        try {
            $computerSystem = Get-WmiObject -Class Win32_ComputerSystem -ErrorAction Stop
            $pageFiles = @(Get-WmiObject -Class Win32_PageFileUsage -ErrorAction Stop)

            $allocatedMiB = 0
            $currentUsageMiB = 0
            $peakUsageMiB = 0
            foreach ($pageFile in $pageFiles) {
                $allocatedMiB += [int]$pageFile.AllocatedBaseSize
                $currentUsageMiB += [int]$pageFile.CurrentUsage
                $peakUsageMiB += [int]$pageFile.PeakUsage
            }

            return [pscustomobject]@{
                AutomaticManaged = [bool]$computerSystem.AutomaticManagedPagefile
                FileCount        = $pageFiles.Count
                AllocatedMiB     = $allocatedMiB
                CurrentUsageMiB  = $currentUsageMiB
                PeakUsageMiB     = $peakUsageMiB
            }
        }
        catch {
            return $null
        }
    }
}

function Test-RcloneLogForOutOfMemory {
    param ([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return $false
    }

    try {
        # Do not inspect only the last few hundred lines. A Go OOM writes the OOM
        # message first and then may dump thousands of goroutine stack lines, which
        # pushes the actual "out of memory" signature far away from the file tail.
        $patterns = @(
            'fatal error:\s*out of memory',
            'out of memory',
            'VirtualAlloc.*failed',
            'errno=1455',
            'paging file is too small'
        )

        foreach ($pattern in $patterns) {
            if (Select-String -LiteralPath $Path -Pattern $pattern -Quiet -ErrorAction Stop) {
                return $true
            }
        }

        return $false
    }
    catch {
        return $false
    }
}

function Invoke-RcloneCopy {
    param (
        [Parameter(Mandatory = $true)][string]$Executable,
        [Parameter(Mandatory = $true)][string]$ConfigPath,
        [Parameter(Mandatory = $true)][string]$RcloneLogPath
    )

    $args = @(
        '--config', $ConfigPath,
        'copy', 'src:', 'dst:',
        '--create-empty-src-dirs',
        '--ignore-existing',
        '--onedrive-no-versions',
        '--onedrive-av-override',
        '--buffer-size', '0',
        '--max-buffer-memory', '64Mi',
        '--max-backlog', [string]$MaxBacklog,
        '--multi-thread-streams', '0',
        '--transfers', [string]$Transfers,
        '--checkers', [string]$Checkers,
        '--onedrive-chunk-size', $ChunkSize,
        '--retries', [string]$RcloneRetries,
        '--retries-sleep', '10s',
        '--low-level-retries', '20',
        '--contimeout', '30s',
        '--timeout', '5m',
        '--stats', '30s',
        '--stats-one-line-date',
        '--stats-log-level', 'INFO',
        '--log-level', 'INFO',
        '--log-file', $RcloneLogPath
    )


    if ($UseChecksum) {
        $args += '--checksum'
    }
    if ($DryRun) {
        $args += '--dry-run'
    }

    $previousErrorActionPreference = $ErrorActionPreference
    $previousGoMemoryLimit = $env:GOMEMLIMIT
    $previousGoGC = $env:GOGC
    $env:GOMEMLIMIT = "${GoMemoryLimitMiB}MiB"
    $env:GOGC = [string]$GoGCPercent
    $ErrorActionPreference = 'Continue'
    try {
        & $Executable @args 2>&1 | ForEach-Object { Write-Host $_ }
        $nativeExitCode = $LASTEXITCODE
    }
    finally {
        $ErrorActionPreference = $previousErrorActionPreference
        if ($null -eq $previousGoMemoryLimit) {
            Remove-Item Env:GOMEMLIMIT -ErrorAction SilentlyContinue
        }
        else {
            $env:GOMEMLIMIT = $previousGoMemoryLimit
        }

        if ($null -eq $previousGoGC) {
            Remove-Item Env:GOGC -ErrorAction SilentlyContinue
        }
        else {
            $env:GOGC = $previousGoGC
        }
    }

    return [int]$nativeExitCode
}

function Invoke-RcloneCleanup {
    param (
        [Parameter(Mandatory = $true)][string]$Executable,
        [Parameter(Mandatory = $true)][string]$ConfigPath,
        [Parameter(Mandatory = $true)][string]$CleanupLogPath,
        [Parameter(Mandatory = $true)][bool]$PreviewOnly
    )

    $args = @(
        '--config', $ConfigPath,
        'cleanup', 'dst:',
        '--checkers', [string]$CleanupCheckers,
        '--stats', '30s',
        '--stats-one-line-date',
        '--stats-log-level', 'INFO',
        '--log-level', 'INFO'
    )

    if ($PreviewOnly) {
        $args += '--dry-run'
    }

    # Do not use --log-file here. Stream each line to the console and append it
    # explicitly as UTF-8 so Windows PowerShell does not create mixed UTF-8/UTF-16 logs.
    # OneDrive cleanup can be slow because rclone must enumerate files and query
    # version history for each one. The 30-second stats prevent a silent console.
    $previousErrorActionPreference = $ErrorActionPreference
    $previousGoMemoryLimit = $env:GOMEMLIMIT
    $previousGoGC = $env:GOGC
    $env:GOMEMLIMIT = "${GoMemoryLimitMiB}MiB"
    $env:GOGC = [string]$GoGCPercent
    $ErrorActionPreference = 'Continue'
    try {
        Add-Content -LiteralPath $CleanupLogPath -Value "===== Cleanup started $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') =====" -Encoding UTF8
        & $Executable @args 2>&1 | ForEach-Object {
            $line = [string]$_
            Write-Host $line
            Add-Content -LiteralPath $CleanupLogPath -Value $line -Encoding UTF8
        }
        $nativeExitCode = $LASTEXITCODE
        Add-Content -LiteralPath $CleanupLogPath -Value "===== Cleanup finished $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') exit=$nativeExitCode =====" -Encoding UTF8
    }
    finally {
        $ErrorActionPreference = $previousErrorActionPreference
        if ($null -eq $previousGoMemoryLimit) {
            Remove-Item Env:GOMEMLIMIT -ErrorAction SilentlyContinue
        }
        else {
            $env:GOMEMLIMIT = $previousGoMemoryLimit
        }

        if ($null -eq $previousGoGC) {
            Remove-Item Env:GOGC -ErrorAction SilentlyContinue
        }
        else {
            $env:GOGC = $previousGoGC
        }
    }

    return [int]$nativeExitCode
}

function Write-RcloneQuota {
    param (
        [Parameter(Mandatory = $true)][string]$Executable,
        [Parameter(Mandatory = $true)][string]$ConfigPath,
        [Parameter(Mandatory = $true)][string]$Label
    )

    Write-Log $Label
    $previousErrorActionPreference = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        & $Executable --config $ConfigPath about 'dst:' 2>&1 | ForEach-Object { Write-Host $_ }
    }
    finally {
        $ErrorActionPreference = $previousErrorActionPreference
    }
}

$results = @()
$hadFailure = $false

try {
    $RclonePath = Resolve-RcloneExecutable -Path $RclonePath
    Test-RcloneClientCredentialsSupport -Executable $RclonePath

    if (-not (Test-Path -LiteralPath $MappingPath -PathType Leaf)) {
        throw "Mapping file not found: $MappingPath"
    }

    $mapping = @(Import-Csv -LiteralPath $MappingPath -Encoding UTF8)
    if ($mapping.Count -eq 0) {
        throw "Mapping file contains no rows: $MappingPath"
    }

    $requiredColumns = @('SourceUPN', 'DestUPN')
    foreach ($column in $requiredColumns) {
        if (-not ($mapping[0].PSObject.Properties.Name -contains $column)) {
            throw "Mapping file is missing required column: $column"
        }
    }

    if ($UserFilter) {
        $mapping = @($mapping | Where-Object { $_.SourceUPN -match $UserFilter })
    }

    if ($mapping.Count -eq 0) {
        Write-Log 'No users matched the supplied filter.' -Type WARN
    }

    Write-Log "Starting migration for $($mapping.Count) user(s). Mapping: $MappingPath"

    # Protect low-memory Windows hosts from spawning several rclone processes at once.
    $detectedMemoryMiB = Get-TotalPhysicalMemoryMiB
    $requestedParallelUsers = $ParallelUsers
    if (-not $InternalWorker -and -not $ForceParallelUsers -and $detectedMemoryMiB) {
        if ($detectedMemoryMiB -lt 2048 -and $ParallelUsers -gt 1) {
            $ParallelUsers = 1
        }
        elseif ($detectedMemoryMiB -lt 4096 -and $ParallelUsers -gt 2) {
            $ParallelUsers = 2
        }

        if ($ParallelUsers -ne $requestedParallelUsers) {
            Write-Log "Low-memory guard: detected approximately $detectedMemoryMiB MiB RAM. Parallel users reduced from $requestedParallelUsers to $ParallelUsers. Use -ForceParallelUsers only if you intentionally accept the OOM risk." -Type WARN
        }
    }

    if ($detectedMemoryMiB -and $detectedMemoryMiB -lt 2048) {
        # Leave an explicit user override alone, otherwise tighten the defaults further
        # for a 1 GiB class Windows VPS.
        if (-not $PSBoundParameters.ContainsKey('GoMemoryLimitMiB') -and $GoMemoryLimitMiB -gt 384) {
            $GoMemoryLimitMiB = 384
        }
        if (-not $PSBoundParameters.ContainsKey('CleanupCheckers') -and $CleanupCheckers -gt 2) {
            $CleanupCheckers = 2
        }

        Write-Log "Low-memory host detected ($detectedMemoryMiB MiB RAM). Using Go memory target ${GoMemoryLimitMiB}MiB and cleanup checkers=$CleanupCheckers." -Type WARN
    }

    $pageFileInfo = Get-PageFileInfo
    if ($pageFileInfo) {
        Write-Log "Windows page file: automaticManaged=$($pageFileInfo.AutomaticManaged), files=$($pageFileInfo.FileCount), allocated=$($pageFileInfo.AllocatedMiB) MiB, currentUsage=$($pageFileInfo.CurrentUsageMiB) MiB, peakUsage=$($pageFileInfo.PeakUsageMiB) MiB."

        if ($pageFileInfo.FileCount -eq 0 -or $pageFileInfo.AllocatedMiB -le 0) {
            Write-Log 'No active Windows page file was detected. This migration has already hit errno=1455; enable a system-managed page file before large-account migrations.' -Type WARN
        }
        elseif ($detectedMemoryMiB -and $detectedMemoryMiB -lt 2048 -and -not $pageFileInfo.AutomaticManaged) {
            Write-Log 'This low-memory host uses a manually sized page file. If errno=1455 recurs, switch the page file to System managed size or increase the commit limit.' -Type WARN
        }
    }
    elseif ($detectedMemoryMiB -and $detectedMemoryMiB -lt 2048) {
        Write-Log 'Unable to query Windows page-file status. Verify that a page file is enabled before migrating the large accounts.' -Type WARN
    }

    $cleanupEnabled = $EnableVersionCleanup.IsPresent
    Write-Log "Options: parallelUsers=$ParallelUsers, transfers=$Transfers, checkers=$Checkers, chunk=$ChunkSize, buffer=0, maxBacklog=$MaxBacklog, goMemoryLimit=${GoMemoryLimitMiB}MiB, goGC=$GoGCPercent, rcloneRetries=$RcloneRetries, wrapperAttempts=$MaxAttempts, ignoreExisting=True, noVersions=True, avOverride=True, checksum=$($UseChecksum.IsPresent), dryRun=$($DryRun.IsPresent), cleanupVersions=$cleanupEnabled, cleanupOnly=$($CleanupOnly.IsPresent), cleanupCheckers=$CleanupCheckers"
    if ($cleanupEnabled -and -not $DryRun) {
        Write-Log 'Version cleanup is ENABLED (-EnableVersionCleanup): after each successful migration, rclone cleanup will PERMANENTLY delete all older versions on that destination OneDrive while keeping the current files. Deleted versions cannot be recovered.' -Type WARN
    }
    elseif (-not $cleanupEnabled) {
        Write-Log 'Version cleanup is DISABLED by default: destination version history is preserved. Pass -EnableVersionCleanup to delete old versions after each successful migration.' -Type WARN
    }

    $parallelHandled = $false

    # Parent process: launch one isolated PowerShell worker per user, with at most
    # $ParallelUsers workers active at once. Each worker runs this same script with
    # -ParallelUsers 1 and -InternalWorker, so rclone configs/logs remain isolated.
    if (-not $InternalWorker -and $ParallelUsers -gt 1 -and $mapping.Count -gt 1) {
        $parallelHandled = $true
        Write-Log "Parallel mode enabled: processing up to $ParallelUsers users simultaneously." -Type SUCCESS

        $childBaseParams = @{
            MappingPath       = $MappingPath
            RclonePath        = $RclonePath
            ParallelUsers     = 1
            InternalWorker    = $true
            Transfers         = $Transfers
            Checkers          = $Checkers
            ChunkSize         = $ChunkSize
            MaxBacklog        = $MaxBacklog
            GoMemoryLimitMiB  = $GoMemoryLimitMiB
            GoGCPercent       = $GoGCPercent
            RcloneRetries     = $RcloneRetries
            MaxAttempts       = $MaxAttempts
            RetryDelaySeconds = $RetryDelaySeconds
            CleanupCheckers   = $CleanupCheckers
        }

        if ($UseChecksum) {
            $childBaseParams['UseChecksum'] = $true
        }
        if ($DryRun) {
            $childBaseParams['DryRun'] = $true
        }
        if ($EnableVersionCleanup) {
            $childBaseParams['EnableVersionCleanup'] = $true
        }
        if ($CleanupOnly) {
            $childBaseParams['CleanupOnly'] = $true
        }

        $pendingUsers = New-Object System.Collections.Queue
        foreach ($parallelUser in $mapping) {
            $pendingUsers.Enqueue([pscustomobject]@{
                SourceUPN = [string]$parallelUser.SourceUPN
                DestUPN   = [string]$parallelUser.DestUPN
            })
        }

        $activeJobs = @()
        $parallelFailure = $false
        $stopScheduling = $false
        $lastParallelHeartbeat = Get-Date

        while ($pendingUsers.Count -gt 0 -or $activeJobs.Count -gt 0) {
            while (-not $stopScheduling -and $pendingUsers.Count -gt 0 -and $activeJobs.Count -lt $ParallelUsers) {
                $parallelUser = $pendingUsers.Dequeue()
                $userFilterExact = '^' + [regex]::Escape($parallelUser.SourceUPN) + '$'

                $childParams = @{}
                foreach ($key in $childBaseParams.Keys) {
                    $childParams[$key] = $childBaseParams[$key]
                }
                $childParams['UserFilter'] = $userFilterExact

                $job = Start-Job -ScriptBlock {
                    param ($ScriptPath, $Parameters, $SourceUPN, $DestUPN)

                    try {
                        $invokeParams = @{}
                        foreach ($key in $Parameters.Keys) {
                            $invokeParams[$key] = $Parameters[$key]
                        }

                        $workerOutput = @(& $ScriptPath @invokeParams)
                        $workerResult = $workerOutput | Where-Object {
                            $_ -and $_.PSObject.Properties.Name -contains 'WorkerExitCode'
                        } | Select-Object -Last 1

                        if ($workerResult) {
                            $workerResult
                        }
                        else {
                            [pscustomobject]@{
                                WorkerExitCode = 1
                                SourceUPN       = $SourceUPN
                                DestUPN         = $DestUPN
                                SummaryPath     = ''
                                LogPath         = ''
                                Message         = 'Worker ended without returning its result object.'
                            }
                        }
                    }
                    catch {
                        [pscustomobject]@{
                            WorkerExitCode = 1
                            SourceUPN       = $SourceUPN
                            DestUPN         = $DestUPN
                            SummaryPath     = ''
                            LogPath         = ''
                            Message         = $_.Exception.Message
                        }
                    }
                } -ArgumentList $PSCommandPath, $childParams, $parallelUser.SourceUPN, $parallelUser.DestUPN

                $activeJobs += [pscustomobject]@{
                    Job       = $job
                    SourceUPN = $parallelUser.SourceUPN
                    DestUPN   = $parallelUser.DestUPN
                }

                Write-Log "[PARALLEL] Started $($parallelUser.SourceUPN) -> $($parallelUser.DestUPN). Active workers: $($activeJobs.Count)/$ParallelUsers"
            }

            if ($activeJobs.Count -eq 0) {
                break
            }

            $finishedJobs = @($activeJobs | Where-Object {
                $_.Job.State -eq 'Completed' -or $_.Job.State -eq 'Failed' -or $_.Job.State -eq 'Stopped'
            })

            if ($finishedJobs.Count -eq 0) {
                $now = Get-Date
                if (($now - $lastParallelHeartbeat).TotalSeconds -ge 30) {
                    $activeNames = (($activeJobs | ForEach-Object { $_.SourceUPN }) -join ', ')
                    Write-Log "[PARALLEL] Still running $($activeJobs.Count) worker(s): $activeNames"
                    $lastParallelHeartbeat = $now
                }
                Start-Sleep -Seconds 2
                continue
            }

            foreach ($finished in $finishedJobs) {
                $jobOutput = @(Receive-Job -Job $finished.Job -ErrorAction SilentlyContinue)
                $workerResult = $jobOutput | Where-Object {
                    $_ -and $_.PSObject.Properties.Name -contains 'WorkerExitCode'
                } | Select-Object -Last 1

                if ($workerResult -and -not [string]::IsNullOrWhiteSpace([string]$workerResult.SummaryPath) -and (Test-Path -LiteralPath $workerResult.SummaryPath)) {
                    $workerRows = @(Import-Csv -LiteralPath $workerResult.SummaryPath -Encoding UTF8)
                    $results += $workerRows
                }

                if (-not $workerResult -or [int]$workerResult.WorkerExitCode -ne 0) {
                    $parallelFailure = $true
                    $hadFailure = $true
                    $workerMessage = 'Unknown worker failure.'
                    if ($workerResult -and -not [string]::IsNullOrWhiteSpace([string]$workerResult.Message)) {
                        $workerMessage = [string]$workerResult.Message
                    }
                    Write-Log "[PARALLEL] Failed: $($finished.SourceUPN) -> $($finished.DestUPN). $workerMessage" -Type ERROR

                    if ($StopOnError) {
                        $stopScheduling = $true
                        Write-Log '[PARALLEL] StopOnError: no additional users will be started. Already-running workers will be allowed to finish.' -Type WARN
                    }
                }
                else {
                    Write-Log "[PARALLEL] Completed: $($finished.SourceUPN) -> $($finished.DestUPN)." -Type SUCCESS
                }

                Remove-Job -Job $finished.Job -Force -ErrorAction SilentlyContinue
                $finishedJobId = $finished.Job.Id
                $activeJobs = @($activeJobs | Where-Object { $_.Job.Id -ne $finishedJobId })
            }
        }

        if ($stopScheduling -and $pendingUsers.Count -gt 0) {
            Write-Log "[PARALLEL] $($pendingUsers.Count) user(s) were not started because StopOnError was triggered." -Type WARN
        }
    }

    if (-not $parallelHandled) {
        $userIndex = 0
        foreach ($user in $mapping) {
        $userIndex++
        $sourceUpn = [string]$user.SourceUPN
        $destUpn = [string]$user.DestUPN
        $sourceId = ''
        if ($user.PSObject.Properties.Name -contains 'SourceId') {
            $sourceId = [string]$user.SourceId
        }

        $destId = ''
        if ($user.PSObject.Properties.Name -contains 'DestId') {
            $destId = [string]$user.DestId
        }

        $sourceUserKey = $sourceUpn
        if (-not [string]::IsNullOrWhiteSpace($sourceId)) {
            $sourceUserKey = $sourceId
        }

        $destUserKey = $destUpn
        if (-not [string]::IsNullOrWhiteSpace($destId)) {
            $destUserKey = $destId
        }

        $startedAt = Get-Date
        $status = 'Failed'
        $exitCode = $null
        $attemptsUsed = 0
        $message = ''
        $configPath = $null
        $rcloneLogPath = ''
        $oomEncountered = $false
        $cleanupStatus = 'NotRun'
        $cleanupExitCode = $null
        $cleanupLogPath = ''

        Write-Log '--------------------------------------------------'
        Write-Log "[$userIndex/$($mapping.Count)] Processing: $sourceUpn -> $destUpn"

        try {
            if ([string]::IsNullOrWhiteSpace($sourceUpn) -or [string]::IsNullOrWhiteSpace($destUpn)) {
                throw 'SourceUPN or DestUPN is empty in the mapping row.'
            }

            Write-Log 'Obtaining short-lived Graph token for destination drive discovery...'
            $dstGraphToken = Get-GraphAccessToken -TenantId $DstTenantId -ClientId $DstClientId -ClientSecret $DstClientSecret

            Write-Log 'Resolving destination OneDrive drive ID...'
            $dstDrive = Get-OneDriveInfo -Token $dstGraphToken -UserId $destUserKey
            Write-Log "Destination drive: $($dstDrive.id) ($($dstDrive.driveType))"

            if ($dstDrive.driveType -ne 'business') {
                throw "Expected OneDrive for Business for $destUpn, but Graph returned driveType '$($dstDrive.driveType)'."
            }

            $safeUser = Get-SafeFileNamePart -Value $sourceUpn
            $safeDestUser = Get-SafeFileNamePart -Value $destUpn
            $tempDir = [System.IO.Path]::GetTempPath()
            if (-not (Test-Path -LiteralPath $tempDir)) {
                New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
            }
            $configPath = Join-Path -Path $tempDir -ChildPath "rclone_migration_${PID}_${safeUser}.conf"
            $rcloneLogBasePath = Join-Path -Path $LogDir -ChildPath "Rclone_${RunTimestamp}_${safeUser}"
            $rcloneLogPath = "${rcloneLogBasePath}_attempt1.log"
            $cleanupLogPath = Join-Path -Path $LogDir -ChildPath "Cleanup_${RunTimestamp}_${safeDestUser}.log"

            if ($CleanupOnly) {
                New-RcloneDestinationConfig -Path $configPath `
                    -DestinationDriveId $dstDrive.id `
                    -DestinationTenantId $DstTenantId `
                    -DestinationClientId $DstClientId `
                    -DestinationClientSecret $DstClientSecret

                if (-not $cleanupEnabled) {
                    throw '-CleanupOnly requires -EnableVersionCleanup: cleanup permanently deletes old destination versions, so it is opt-in.'
                }

                $status = 'CleanupPending'
                $message = 'Copy skipped because -CleanupOnly was specified.'
            }
            else {
                Write-Log 'Obtaining short-lived Graph token for source drive discovery...'
                $srcGraphToken = Get-GraphAccessToken -TenantId $SrcTenantId -ClientId $SrcClientId -ClientSecret $SrcClientSecret

                Write-Log 'Resolving source OneDrive drive ID...'
                $srcDrive = Get-OneDriveInfo -Token $srcGraphToken -UserId $sourceUserKey
                Write-Log "Source drive: $($srcDrive.id) ($($srcDrive.driveType))"

                New-RcloneConfig -Path $configPath `
                    -SourceDriveId $srcDrive.id `
                    -DestinationDriveId $dstDrive.id `
                    -SourceTenantId $SrcTenantId `
                    -SourceClientId $SrcClientId `
                    -SourceClientSecret $SrcClientSecret `
                    -DestinationTenantId $DstTenantId `
                    -DestinationClientId $DstClientId `
                    -DestinationClientSecret $DstClientSecret

                for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
                    $attemptsUsed = $attempt
                    $rcloneLogPath = "${rcloneLogBasePath}_attempt${attempt}.log"
                    Write-Log "Starting rclone copy (attempt $attempt/$MaxAttempts). Detailed log: $rcloneLogPath"

                    $exitCode = Invoke-RcloneCopy -Executable $RclonePath -ConfigPath $configPath -RcloneLogPath $rcloneLogPath

                    if ($exitCode -eq 0) {
                        if ($DryRun) {
                            $status = 'DryRun'
                            $message = 'Copy dry run completed successfully.'
                        }
                        else {
                            $status = 'Success'
                            $message = 'Migration completed successfully.'
                        }
                        Write-Log "$message $sourceUpn -> $destUpn" -Type SUCCESS
                        break
                    }

                    Write-Log "rclone exited with code $exitCode for $sourceUpn." -Type WARN

                    # Exit code 5 is rclone's documented temporary/retryable error.
                    # Exit code 2 normally means syntax/usage error and is NOT blindly retried.
                    # The Go runtime may also terminate with code 2 after a Windows OOM, so in
                    # that specific case we detect the OOM signature in the rclone log and retry.
                    $oomDetected = Test-RcloneLogForOutOfMemory -Path $rcloneLogPath
                    $retryable = ($exitCode -eq 5)
                    $retryReason = 'temporary rclone error'

                    if ($oomDetected) {
                        $oomEncountered = $true
                        $retryable = $true
                        $retryReason = 'out-of-memory termination'

                        $oldTransfers = $Transfers
                        $oldCheckers = $Checkers
                        $oldGoMemoryLimitMiB = $GoMemoryLimitMiB
                        $oldMaxBacklog = $MaxBacklog

                        $Transfers = [Math]::Max(1, [int][Math]::Floor($Transfers / 2))
                        $Checkers = [Math]::Max(1, [int][Math]::Floor($Checkers / 2))
                        $GoMemoryLimitMiB = [Math]::Max(192, [int][Math]::Floor($GoMemoryLimitMiB * 0.75))
                        $MaxBacklog = [Math]::Max(250, [int][Math]::Floor($MaxBacklog / 2))

                        Write-Log "OOM detected in current attempt log. Reducing this worker from transfers=$oldTransfers/checkers=$oldCheckers/goMemory=${oldGoMemoryLimitMiB}MiB/maxBacklog=$oldMaxBacklog to transfers=$Transfers/checkers=$Checkers/goMemory=${GoMemoryLimitMiB}MiB/maxBacklog=$MaxBacklog before retry." -Type WARN
                    }

                    if (-not $retryable -or $attempt -ge $MaxAttempts) {
                        if ($oomDetected) {
                            $message = "rclone terminated because Windows ran out of committed memory. Exit code $exitCode. See $rcloneLogPath"
                        }
                        else {
                            $message = "rclone failed with exit code $exitCode. See $rcloneLogPath"
                        }
                        break
                    }

                    $delay = [Math]::Min(300, $RetryDelaySeconds * [Math]::Pow(2, $attempt - 1))
                    Write-Log "Retryable $retryReason. Retrying in $delay seconds..." -Type WARN
                    Start-Sleep -Seconds $delay
                }

                if ($status -eq 'Failed' -and [string]::IsNullOrWhiteSpace($message)) {
                    $message = "rclone did not complete successfully. See $rcloneLogPath"
                }
            }

            $canCleanup = $false
            if ($CleanupOnly) {
                $canCleanup = $true
            }
            elseif ($status -eq 'Success' -or $status -eq 'DryRun') {
                $canCleanup = $true
            }

            if ($cleanupEnabled -and $canCleanup) {
                if ($DryRun) {
                    Write-Log "Previewing old-version cleanup for $destUpn. No versions will be deleted. Detailed log: $cleanupLogPath" -Type WARN
                    $cleanupStatus = 'DryRun'
                }
                else {
                    Write-Log "Cleaning ALL old OneDrive versions for $destUpn. Current files are kept. Detailed log: $cleanupLogPath" -Type WARN
                    Write-RcloneQuota -Executable $RclonePath -ConfigPath $configPath -Label "Destination quota before cleanup for ${destUpn}:"
                    $cleanupStatus = 'Running'
                }

                $cleanupExitCode = Invoke-RcloneCleanup -Executable $RclonePath -ConfigPath $configPath -CleanupLogPath $cleanupLogPath -PreviewOnly $DryRun.IsPresent

                if ($cleanupExitCode -ne 0) {
                    $cleanupStatus = 'Failed'
                    if ($CleanupOnly) {
                        $status = 'Failed'
                    }
                    else {
                        $status = 'CleanupFailed'
                    }
                    $message = "Version cleanup failed with exit code $cleanupExitCode. Migration current files were kept. See $cleanupLogPath"
                    Write-Log $message -Type ERROR
                }
                else {
                    if ($DryRun) {
                        $cleanupStatus = 'DryRunSuccess'
                        if ($CleanupOnly) {
                            $status = 'CleanupDryRun'
                        }
                        $message = 'Cleanup dry run completed successfully; no versions were deleted.'
                        Write-Log "$message $destUpn" -Type SUCCESS
                    }
                    else {
                        $cleanupStatus = 'Success'
                        if ($CleanupOnly) {
                            $status = 'CleanupSuccess'
                            $message = 'Old-version cleanup completed successfully.'
                        }
                        else {
                            $message = 'Migration and old-version cleanup completed successfully.'
                        }
                        Write-RcloneQuota -Executable $RclonePath -ConfigPath $configPath -Label "Destination quota after cleanup for ${destUpn}:"
                        Write-Log "$message $destUpn" -Type SUCCESS
                    }
                }
            }
            elseif (-not $cleanupEnabled) {
                $cleanupStatus = 'Skipped'
            }
            elseif (-not $canCleanup) {
                $cleanupStatus = 'SkippedDueToMigrationFailure'
            }
        }
        catch {
            $message = $_.Exception.Message
            Write-Log "Migration failed for $sourceUpn -> $destUpn. Error: $message" -Type ERROR
        }
        finally {
            if ($configPath -and (Test-Path -LiteralPath $configPath)) {
                Remove-Item -LiteralPath $configPath -Force -ErrorAction SilentlyContinue
            }

            $endedAt = Get-Date
            $duration = $endedAt - $startedAt
            $results += [pscustomobject]@{
                SourceUPN       = $sourceUpn
                DestUPN         = $destUpn
                Status          = $status
                ExitCode        = $exitCode
                Attempts        = $attemptsUsed
                StartedAt       = $startedAt.ToString('yyyy-MM-dd HH:mm:ss')
                EndedAt         = $endedAt.ToString('yyyy-MM-dd HH:mm:ss')
                DurationMinutes = [Math]::Round($duration.TotalMinutes, 2)
                RcloneLog       = $rcloneLogPath
                OOMEncountered  = $oomEncountered
                EffectiveTransfers = $Transfers
                EffectiveCheckers  = $Checkers
                EffectiveGoMemoryLimitMiB = $GoMemoryLimitMiB
                CleanupStatus   = $cleanupStatus
                CleanupExitCode = $cleanupExitCode
                CleanupLog      = $cleanupLogPath
                Message         = $message
            }
        }

        if ($status -eq 'Failed' -or $status -eq 'CleanupFailed') {
            $hadFailure = $true
            if ($StopOnError) {
                Write-Log 'StopOnError was requested; stopping the batch.' -Type ERROR
                break
            }
        }
        }
    }
}
catch {
    $hadFailure = $true
    Write-Log "Fatal orchestrator error: $($_.Exception.Message)" -Type ERROR
}
finally {
    if ($results.Count -gt 0) {
        $results | Export-Csv -LiteralPath $SummaryPath -NoTypeInformation -Encoding UTF8
        Write-Log "Summary written to: $SummaryPath"
    }

    # Best-effort removal of plaintext secrets from script variables.
    $SrcClientSecret = $null
    $DstClientSecret = $null
    Remove-Variable srcGraphToken, dstGraphToken -ErrorAction SilentlyContinue
}

$finalExitCode = 0
if ($hadFailure) {
    $finalExitCode = 1
    Write-Log 'Migration orchestrator completed with one or more failures.' -Type ERROR
}
elseif ($CleanupOnly) {
    Write-Log 'Version-cleanup batch completed successfully.' -Type SUCCESS
}
else {
    Write-Log 'Migration orchestrator completed successfully.' -Type SUCCESS
}

if ($InternalWorker) {
    $workerSourceUpn = ''
    $workerDestUpn = ''
    $mappingVariable = Get-Variable -Name mapping -ErrorAction SilentlyContinue
    if ($mappingVariable -and $mapping -and $mapping.Count -gt 0) {
        $workerSourceUpn = [string]$mapping[0].SourceUPN
        $workerDestUpn = [string]$mapping[0].DestUPN
    }

    $workerMessage = 'Worker completed successfully.'
    if ($finalExitCode -ne 0) {
        $workerMessage = 'Worker completed with one or more failures.'
    }

    [pscustomobject]@{
        WorkerExitCode = $finalExitCode
        SourceUPN       = $workerSourceUpn
        DestUPN         = $workerDestUpn
        SummaryPath     = $SummaryPath
        LogPath         = $LogPath
        Message         = $workerMessage
    }
    return
}

exit $finalExitCode
