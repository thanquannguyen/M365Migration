<#
.SYNOPSIS
Uploads CHU_Y.txt to standard OneDrive folders on each source account.

Requires (minimum Graph scopes): User.Read.All, Files.ReadWrite.All,
Sites.ReadWrite.All, Directory.Read.All.
#>

[CmdletBinding()]
param (
    [string]$ConfigPath = '.\MigrationConfig.psd1',
    [string]$MappingPath = '.\Output\UserMapping.csv',
    [string]$UserFilter
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
$LogPath = Join-Path -Path $LogDir -ChildPath "Announcement_${Timestamp}${LogSuffix}_${PID}.log"

function Write-Log {
    param ([string]$Message, [string]$Type = "INFO")
    $LogMsg = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$Type] $Message"
    Write-Host $LogMsg -ForegroundColor $(if ($Type -eq 'ERROR') {'Red'} elseif ($Type -eq 'WARN') {'Yellow'} else {'Cyan'})
    Add-Content -Path $LogPath -Value $LogMsg -Encoding UTF8 -ErrorAction SilentlyContinue
}

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
$Mapping = @(Import-Csv -Path $MappingPath -Encoding UTF8)
$isFilter = [bool]$UserFilter
if ($isFilter) {
    # S16: exact-match first (UPN or bare username); unescaped -match only
    # as a fallback when nothing matched exactly.
    $exactMatches = @($Mapping | Where-Object { $_.SourceUPN -eq $UserFilter -or $_.SourceUPN -like "$UserFilter@*" -or $_.DestUPN -eq $UserFilter -or $_.DestUPN -like "$UserFilter@*" })
    if ($exactMatches.Count -gt 0) {
        $Mapping = $exactMatches
    } else {
        $Mapping = @($Mapping | Where-Object { $_.SourceUPN -match $UserFilter -or $_.DestUPN -match $UserFilter })
    }
    Write-Log "Filtered to user(s): $UserFilter ($($Mapping.Count) user(s))"
}

$MaxConcurrentUsers = if ($null -ne $config.MaxConcurrentUsers -and $config.MaxConcurrentUsers -gt 0) { [int]$config.MaxConcurrentUsers } else { 4 }
$MaxConcurrentUsers = [Math]::Min($MaxConcurrentUsers, $Mapping.Count)

if (-not $isFilter -and $Mapping.Count -gt 0) {
    Write-Host "`n==========================================================" -ForegroundColor Cyan
    Write-Host "   PARALLEL ANNOUNCEMENT ORCHESTRATOR" -ForegroundColor Cyan
    Write-Host "   Concurrent Threads: $MaxConcurrentUsers | Total Accounts: $($Mapping.Count)" -ForegroundColor Cyan
    Write-Host "==========================================================`n" -ForegroundColor Cyan
    
    $pendingUsers = [System.Collections.ArrayList]@($Mapping)
    $runningJobs = @{}
    $retryCounts = @{}
    $scriptPath = $MyInvocation.MyCommand.Definition
    
    Write-Host "[STEP 1] Authenticating with Source Tenant before launching parallel threads..." -ForegroundColor Yellow
    $sourceTenantId = $config.SourceTenantDomain
    $sourceScopes = @('User.Read.All', 'Files.ReadWrite.All', 'Sites.ReadWrite.All', 'Directory.Read.All')
    Connect-MgGraph -Scopes $sourceScopes -NoWelcome -ErrorAction SilentlyContinue
    if (-not (Get-MgContext)) {
        if ($sourceTenantId) {
            Connect-MgGraph -Scopes $sourceScopes -TenantId $sourceTenantId -UseDeviceCode -NoWelcome *>&1 | Write-Host
        } else {
            Connect-MgGraph -Scopes $sourceScopes -UseDeviceCode -NoWelcome *>&1 | Write-Host
        }
    }
    Write-Host "Connected as $((Get-MgContext).Account). Launching parallel threads...`n" -ForegroundColor Green

    try {
        while ($pendingUsers.Count -gt 0 -or $runningJobs.Count -gt 0) {
            while ($runningJobs.Count -lt $MaxConcurrentUsers -and $pendingUsers.Count -gt 0) {
                $user = $pendingUsers[0]
                $pendingUsers.RemoveAt(0)
                $upn = $user.SourceUPN
                Write-Host "[LAUNCHING] Thread for account: $upn" -ForegroundColor Green
                $job = Start-Job -Name "ANN-$upn" -ScriptBlock {
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
                        Write-Host "[COMPLETED] Account: $key" -ForegroundColor Green
                    } else {
                        Write-Host "[FAILED/STOPPED] Account: $key - State: $($j.State)" -ForegroundColor Yellow
                        $output | ForEach-Object { Write-Host "   $_" -ForegroundColor DarkRed }
                        # S2: retry cap (3 attempts total) with backoff; a permanently
                        # failing account is recorded as failed instead of re-queued forever.
                        $retryCounts[$key] = [int]$retryCounts[$key] + 1
                        if ($retryCounts[$key] -lt 3) {
                            Write-Host "[RETRY $($retryCounts[$key])/3] Account: $key - Re-queueing after $(5 * $retryCounts[$key])s backoff..." -ForegroundColor Yellow
                            Start-Sleep -Seconds (5 * $retryCounts[$key])
                            $mappingRow = $Mapping | Where-Object { $_.SourceUPN -eq $key } | Select-Object -First 1
                            if ($mappingRow) { [void]$pendingUsers.Add($mappingRow) }
                        } else {
                            Write-Host "[FAILED PERMANENTLY] Account: $key - failed 3 times, not re-queued." -ForegroundColor Red
                        }
                    }
                    Remove-Job -Job $j -Force -ErrorAction SilentlyContinue | Out-Null
                }
            }
            foreach ($c in $completed) {
                $runningJobs.Remove($c)
            }
            
            if ($runningJobs.Count -gt 0) {
                Write-Host "`n--- LIVE ANNOUNCEMENT PROGRESS ($($runningJobs.Count) accounts active) ---" -ForegroundColor Cyan
                Start-Sleep -Seconds 5
            }
        }
        Write-Host "`n[ALL ACCOUNTS ANNOUNCEMENT FINISHED]" -ForegroundColor Green
    } finally {
        foreach ($job in $runningJobs.Values) {
            Stop-Job -Job $job -ErrorAction SilentlyContinue
            Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
        }
        Get-Job | Where-Object { $_.Name -like "ANN-*" } | ForEach-Object { Stop-Job -Job $_ -ErrorAction SilentlyContinue; Remove-Job -Job $_ -Force -ErrorAction SilentlyContinue }
    }
    return
}

# ----------------- SINGLE USER MODE -----------------
$sourceTenantId = $config.SourceTenantDomain
$sourceScopes = @('User.Read.All', 'Files.ReadWrite.All', 'Sites.ReadWrite.All', 'Directory.Read.All')

try { Set-MgGraphOption -DisableLoginByWAM $true -ErrorAction SilentlyContinue } catch {}

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
        Write-Log "Background job failed to acquire token. Aborting." -Type ERROR
        return
    }
    if ($config.SourceClientId -and $config.SourceClientSecret) {
        $secSecret = ConvertTo-SecureString $config.SourceClientSecret -AsPlainText -Force
        $cred = New-Object System.Management.Automation.PSCredential($config.SourceClientId, $secSecret)
        Connect-MgGraph -ClientSecretCredential $cred -TenantId $sourceTenantId -NoWelcome *>$null
    } else {
        Write-Log "A device code will appear. Sign in with SOURCE tenant admin." -Type WARN
        if ($sourceTenantId) {
            Connect-MgGraph -Scopes $sourceScopes -TenantId $sourceTenantId -UseDeviceCode -NoWelcome *>&1 | Write-Host
        } else {
            Connect-MgGraph -Scopes $sourceScopes -UseDeviceCode -NoWelcome *>&1 | Write-Host
        }
    }
}

$Summary = @()

foreach ($User in $Mapping) {
    $SourceUpn = $User.SourceUPN
    $SourceId = $User.SourceId
    $DestUpn = $User.DestUPN

    Write-Log "Processing announcement for $SourceUpn"

    Write-Log "  Getting DriveId for $SourceUpn..."
    $driveId = Get-UserOneDriveId -UserId $SourceId
    if (-not $driveId) {
        Write-Log "Cannot access OneDrive for $SourceUpn" -Type WARN
        $Summary += [PSCustomObject]@{ SourceUpn = $SourceUpn; Status = 'Failed - No Drive' }
        continue
    }
    Write-Log "  Got DriveId: $driveId"

    # B7/S19: text + upload live in AnnouncementHelpers.ps1. The template
    # contains NO password; users are told to contact their administrator.
    $content = Get-AnnouncementText -SourceUpn $SourceUpn -DestUpn $DestUpn
    $uploadedCount = Publish-AnnouncementToFolders -SourceUpn $SourceUpn -SourceId $SourceId -DriveId $driveId -Content $content -LogAction { param($m, $t) Write-Log $m -Type $t }
    Write-Log "  Announcement uploaded to $uploadedCount folder(s) for $SourceUpn"

    $Summary += [PSCustomObject]@{ SourceUpn = $SourceUpn; Status = "Success ($uploadedCount files)" }
}

$Summary | Format-Table -AutoSize | Out-String | Write-Host -ForegroundColor Cyan
Write-Log "Announcement script completed."
